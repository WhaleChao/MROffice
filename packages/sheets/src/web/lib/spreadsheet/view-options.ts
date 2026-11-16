/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * Workbook-level VIEW options (Excel / WPS parity) and their pure helpers, kept
 * out of the React component so they unit-test in isolation and round-trip
 * predictably. These are DISPLAY-only settings persisted at the spreadsheet
 * content top level (alongside `referenceStyle`); they never change the stored
 * cells, so a default workbook serializes byte-for-byte as before.
 *
 *  - Show Formulas (Ctrl+`) and the sheet's view (Normal / Page Layout / Page
 *    Break Preview) are not here either: Excel keeps them on each
 *    worksheet's view too (`showFormulas` / `viewMode` on a sheet), and a
 *    page that stored them here is read onto every sheet by normalizeSheets.
 *  - View > Gridlines and View > Headings are not here: Excel keeps them on
 *    each worksheet's view, and so does the sheet model (`hideGridlines` /
 *    `hideHeadings` on a sheet). A page that stored them here, for the whole
 *    workbook, is read onto every sheet by normalizeSheets.
 *  - hideFormulaBar: hide the formula bar (View > Formula Bar off). Excel has
 *    had this since 2007; a sheet used as a dashboard wants the strip gone.
 *  - customLists: the user's saved Custom Lists for drag-fill (File > Options >
 *    Custom Lists), each a series of tokens.
 *  - readingMode: Excel "Focus Cell" (View ▸ Show; WPS calls it 阅读模式) —
 *    the active cell's whole row and column are tinted, so the eye can follow
 *    a wide table across the screen without losing its line.
 *  - focusCellColor: Excel's "Focus Cell Color" — the colour that tints the
 *    crosshair, picked from the editors' shared palette. Excel's default is
 *    Green, and it paints the tint lighter than the swatch it shows. Excel
 *    keeps the choice as an application setting that outlives the toggle;
 *    here it rides with the other view options so it survives a reload the
 *    same way.
 *  - eyeCare: WPS "护眼模式" — the sheet paper turns a low-saturation green
 *    instead of white, the setting people leave on for long sessions.
 *  - spellCheck: whether the cell editor asks the browser to spell-check what
 *    is being typed. On by default (the browser's own default for a text
 *    input); a sheet full of part numbers and codes wants it off, which is
 *    the only reason the setting exists.
 *  - themeId: the workbook's colour theme (Page Layout ▸ Themes). It selects
 *    the palette and repaints semantic theme bindings such as built-in table
 *    styles, charts and future shapes. A literal colour stored directly on a
 *    cell remains direct formatting and therefore does not change.
 *
 * The booleans are stored as the NON-default value only (Excel's default is
 * the formula bar shown), so the persisted JSON stays minimal
 * and a never-touched workbook is unchanged.
 */

import { matchingChartValueCache } from "@mrnote/core/web/lib/spreadsheet/chart-value-cache.ts";
import {
  normalizeChartTemplates,
  type SpreadsheetChartTemplate,
} from "./chart-templates";
import {
  normalizeNativeTableStyles,
  type SpreadsheetNativeTableStyle,
} from "@mrnote/core/web/lib/spreadsheet/native-table-styles.ts";
import { normalizeHexColor } from "@mrnote/core/web/lib/office-color-palette.ts";
import {
  normalizeCustomViews,
  type SpreadsheetCustomView,
} from "./custom-views";
import {
  normalizeCustomCellStyles,
  type CustomCellStyle,
} from "./cell-styles";
import {
  isSpreadsheetThemePaletteId,
  paletteById,
  type SpreadsheetNativeThemeColors,
} from "@mrnote/core/web/lib/spreadsheet/color-palette.ts";
import {
  normalizeSpreadsheetNativeThemeColors,
  normalizeSpreadsheetNativeThemeEffects,
  normalizeSpreadsheetNativeThemeFonts,
  normalizeSpreadsheetThemeEffectId,
  normalizeSpreadsheetThemeFontId,
  type SpreadsheetNativeThemeEffects,
  type SpreadsheetNativeThemeFonts,
  type SpreadsheetThemeFontSetId,
  type SpreadsheetThemeEffectId,
} from "@mrnote/core/web/lib/spreadsheet/workbook-theme.ts";

export type SpreadsheetViewMode =
  | "normal"
  | "pageLayout"
  | "pageBreakPreview";

/**
 * Excel highlights in Green until the user picks another colour. The colour
 * is chosen from the editors' one shared palette (theme colours, standard
 * colours, recents, "More colours…"), so this is that palette's standard
 * green rather than Excel's own swatch value.
 */
export const DEFAULT_FOCUS_CELL_COLOR = "#16a34a";

/** A stored focus-cell colour as `#rrggbb` (the default when absent or junk). */
export function resolveFocusCellColor(raw: unknown): string {
  return normalizeHexColor(raw) ?? DEFAULT_FOCUS_CELL_COLOR;
}

/** The colour to persist: only a valid, non-default `#rrggbb`. */
export function normalizeFocusCellColor(raw: unknown): string | undefined {
  const hex = normalizeHexColor(raw);
  return hex && hex !== DEFAULT_FOCUS_CELL_COLOR ? hex : undefined;
}

/**
 * The tint's colour channels as a CSS `R G B` triple, for the overlay to mix
 * with its own alpha (the tint Excel paints is lighter than the swatch, so the
 * swatch colour is never painted at full strength).
 */
export function focusCellTintChannels(raw: string | undefined): string {
  const hex = resolveFocusCellColor(raw);
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `${r} ${g} ${b}`;
}

/** The normalized workbook view settings. */
export interface SpreadsheetViewOptions {
  hideFormulaBar: boolean;
  /** User custom fill lists; each a list of >= 2 trimmed, non-empty tokens. */
  customLists: string[][];
  readingMode: boolean;
  /** Focus Cell tint colour as `#rrggbb`; absent means the default green. */
  focusCellColor?: string;
  eyeCare: boolean;
  themeId: string;
  /** Workbook theme font set; direct cell font formatting still wins. */
  themeFontSetId?: SpreadsheetThemeFontSetId;
  /** Workbook theme drawing effects consumed by newly inserted shapes. */
  themeEffectId?: SpreadsheetThemeEffectId;
  /**
   * The workbook's own color, font and effect schemes (ones Office does not
   * ship), as its file has them. Each replaces the matching id above and
   * stays until a gallery choice replaces that scheme.
   */
  nativeThemeColors?: SpreadsheetNativeThemeColors;
  nativeThemeFonts?: SpreadsheetNativeThemeFonts;
  nativeThemeEffects?: SpreadsheetNativeThemeEffects;
  /** A theme MRAI cannot read stays package-owned: nothing may change it. */
  nativeThemeDisposition?: "preserve-only";
  /** Read-only source accents. They never authorize writing a custom theme. */
  nativeThemeAccents?: string[];
  /**
   * The unreadable theme sits in the one part the workbook names for its
   * theme, so applying a whole theme (Page Layout > Themes) may replace it,
   * as Excel does. Its single schemes still cannot change.
   */
  nativeThemeReplaceable?: true;
  /**
   * The workbook's own table styles (the gallery's Custom section), as its
   * styles part defines them. Read-only package data: tables paint and choose
   * them, nothing edits the definitions.
   */
  nativeTableStyles?: SpreadsheetNativeTableStyle[];
  /** The workbook's own cell styles (New Cell Style… / Merge Styles…,
   * XLS-A012): the Cell Styles gallery's Custom section. */
  customCellStyles?: CustomCellStyle[];
  /**
   * Formulas ▸ Calculation Options ▸ Automatic Except for Data Tables
   * (XLS-A025, xlsx calcMode="autoNoTable"); meaningful with Automatic only.
   */
  calculationExceptDataTables?: true;
  /**
   * File ▸ Options ▸ Formulas ▸ Enable iterative calculation, Maximum
   * Iterations, Maximum Change (xlsx calcPr@iterate, @iterateCount,
   * @iterateDelta). Kept when off too, as Excel keeps the limits.
   */
  calculationIteration?: SpreadsheetCalculationIteration;
  /**
   * File ▸ Options ▸ Advanced ▸ Set precision as displayed (xlsx
   * calcPr@fullPrecision="0"): a formula's number is kept as its format shows
   * it, and the formulas that read it read that.
   */
  precisionAsDisplayed?: true;
  /** View ▸ Show ▸ Ruler off (Page Layout view's rulers, XLS-A030). */
  hideRuler?: true;
  /** View ▸ Workbook Views ▸ Custom Views (XLS-A030). */
  customViews?: SpreadsheetCustomView[];
  /** XLS-A081 Save as Template: the workbook's Templates folder of charts. */
  chartTemplates?: SpreadsheetChartTemplate[];
  /**
   * XLS-A056: the number formats the workbook defines (its file's
   * `numFmts`), the ones no cell uses included: Format Cells > Custom lists
   * them and Delete removes one. Absent on a page read before they were read.
   */
  numberFormats?: string[];
  spellCheck: boolean;
}

/** The all-default view options (everything Excel-default). */
export function defaultViewOptions(): SpreadsheetViewOptions {
  return {
    hideFormulaBar: false,
    customLists: [],
    readingMode: false,
    eyeCare: false,
    themeId: "office",
    spellCheck: true,
  };
}

/** Hard cap on custom lists / members so a hand-edited file can't blow up. */
const MAX_CUSTOM_LISTS = 64;
const MAX_LIST_MEMBERS = 256;
const MAX_MEMBER_LENGTH = 256;

/**
 * Normalize a single persisted custom list: keep string members, trim them,
 * drop empties, bound their length, de-duplicate case-insensitively, and cap the
 * count. Returns null when fewer than two distinct members survive (a one-item
 * list is just a constant and is meaningless as a fill series).
 */
function normalizeList(raw: unknown): string[] | null {
  if (!Array.isArray(raw)) return null;
  const out: string[] = [];
  const seen = new Set<string>();
  for (const item of raw) {
    if (out.length >= MAX_LIST_MEMBERS) break;
    if (typeof item !== "string") continue;
    const trimmed = item.trim().slice(0, MAX_MEMBER_LENGTH);
    if (trimmed === "") continue;
    const lower = trimmed.toLowerCase();
    if (seen.has(lower)) continue;
    seen.add(lower);
    out.push(trimmed);
  }
  return out.length >= 2 ? out : null;
}

/**
 * Normalize the persisted custom-lists value (an array of arrays). Drops invalid
 * lists, caps the total count, and returns a fresh array of clean lists. Safe on
 * any input (non-array, nested junk) — returns [] when nothing valid survives.
 */
export function normalizeCustomLists(raw: unknown): string[][] {
  if (!Array.isArray(raw)) return [];
  const out: string[][] = [];
  for (const entry of raw) {
    if (out.length >= MAX_CUSTOM_LISTS) break;
    const list = normalizeList(entry);
    if (list) out.push(list);
  }
  return out;
}

/**
 * Normalize an unknown persisted view-options blob (read from the spreadsheet
 * content top level) into a complete {@link SpreadsheetViewOptions}. Unknown /
 * missing fields fall back to the Excel default. Booleans accept only literal
 * `true` (so a stray truthy value can't silently flip a default).
 */
/** Enable iterative calculation and its limits (Excel's defaults 100, 0.001). */
export type SpreadsheetCalculationIteration = {
  enabled: boolean;
  maxIterations: number;
  maxChange: number;
};
export const DEFAULT_CALCULATION_ITERATION: SpreadsheetCalculationIteration = {
  enabled: false,
  maxIterations: 100,
  maxChange: 0.001,
};

/** A stored iteration setting, or null when it states none (or a malformed one). */
export function normalizeCalculationIteration(raw: unknown): SpreadsheetCalculationIteration | null {
  if (typeof raw !== "object" || raw === null) return null;
  const { enabled, maxIterations, maxChange } = raw as Record<string, unknown>;
  if (typeof enabled !== "boolean") return null;
  // Excel takes 1..32767 passes and a change of at least 0.
  if (typeof maxIterations !== "number" || !Number.isInteger(maxIterations) || maxIterations < 1 || maxIterations > 32767) return null;
  if (typeof maxChange !== "number" || !Number.isFinite(maxChange) || maxChange < 0) return null;
  return { enabled, maxIterations, maxChange };
}

/**
 * The calculation session's `iteration` for a workbook's setting (off:
 * none), with each formula cell's current number as the seed its first
 * iterations start from (the cell's stored formula cache: the file's cached
 * value, or the last save's), as Excel starts from the values cells hold.
 */
export function calculationIterationOption(
  options: Pick<SpreadsheetViewOptions, "calculationIteration">,
  sheets?: readonly { id: string; name?: unknown; cells: Readonly<Record<string, { value?: unknown; chartValueCache?: unknown } | undefined>> }[],
): { maxIterations: number; maxChange: number; seeds?: Array<[string, number, number, number]> } | undefined {
  const setting = options.calculationIteration;
  if (!setting?.enabled) return undefined;
  const seeds: Array<[string, number, number, number]> = [];
  for (const sheet of sheets ?? []) {
    const name = typeof sheet.name === "string" && sheet.name ? sheet.name : sheet.id;
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (typeof cell?.value !== "string" || !matchingChartValueCache(cell.chartValueCache, cell.value)) continue;
      const [row, col] = key.split(":").map(Number);
      if (Number.isSafeInteger(row) && Number.isSafeInteger(col)) seeds.push([name, row, col, cell.chartValueCache.value]);
    }
  }
  return {
    maxIterations: setting.maxIterations,
    maxChange: setting.maxChange,
    ...(seeds.length ? { seeds } : {}),
  };
}

export function normalizeViewOptions(raw: unknown): SpreadsheetViewOptions {
  const source =
    raw && typeof raw === "object" ? (raw as Record<string, unknown>) : {};
  const calculationIteration = normalizeCalculationIteration(source.calculationIteration);
  const themeFontSetId = normalizeSpreadsheetThemeFontId(
    source.themeFontSetId,
  );
  const themeEffectId = normalizeSpreadsheetThemeEffectId(
    source.themeEffectId,
  );
  const preserveOnly = source.nativeThemeDisposition === "preserve-only";
  const focusCellColor = normalizeFocusCellColor(source.focusCellColor);
  const nativeThemeColors = preserveOnly
    ? null
    : normalizeSpreadsheetNativeThemeColors(source.nativeThemeColors);
  const nativeThemeFonts = preserveOnly
    ? null
    : normalizeSpreadsheetNativeThemeFonts(source.nativeThemeFonts);
  const nativeThemeEffects = preserveOnly
    ? null
    : normalizeSpreadsheetNativeThemeEffects(source.nativeThemeEffects);
  const nativeTableStyles = normalizeNativeTableStyles(source.nativeTableStyles);
  const customCellStyles = normalizeCustomCellStyles(source.customCellStyles);
  const customViews = normalizeCustomViews(source.customViews);
  const chartTemplates = normalizeChartTemplates(source.chartTemplates);
  const numberFormats = Array.isArray(source.numberFormats)
    ? [...new Set(source.numberFormats.filter(
        (code): code is string => typeof code === "string" && code.length > 0 && code.length <= 255,
      ))].slice(0, 250)
    : undefined;
  const nativeThemeAccents = source.nativeThemeDisposition === "preserve-only" &&
    Array.isArray(source.nativeThemeAccents) && source.nativeThemeAccents.length === 6 &&
    source.nativeThemeAccents.every((color) => typeof color === "string" && /^#[0-9a-f]{6}$/i.test(color))
      ? source.nativeThemeAccents.map((color: string) => color.toLowerCase()) : undefined;
  return {
    hideFormulaBar: source.hideFormulaBar === true,
    customLists: normalizeCustomLists(source.customLists),
    readingMode: source.readingMode === true,
    ...(focusCellColor ? { focusCellColor } : {}),
    eyeCare: source.eyeCare === true,
    themeId:
      !nativeThemeColors && isSpreadsheetThemePaletteId(source.themeId)
        ? source.themeId
        : "office",
    ...(nativeThemeFonts || themeFontSetId === "office" ? {} : { themeFontSetId }),
    ...(nativeThemeEffects || themeEffectId === "none" ? {} : { themeEffectId }),
    ...(nativeThemeColors ? { nativeThemeColors } : {}),
    ...(nativeThemeFonts ? { nativeThemeFonts } : {}),
    ...(nativeThemeEffects ? { nativeThemeEffects } : {}),
    ...(source.nativeThemeDisposition === "preserve-only"
      ? { nativeThemeDisposition: "preserve-only" as const }
      : {}),
    ...(nativeThemeAccents ? { nativeThemeAccents } : {}),
    ...(preserveOnly && source.nativeThemeReplaceable === true
      ? { nativeThemeReplaceable: true as const }
      : {}),
    ...(nativeTableStyles.length > 0 ? { nativeTableStyles } : {}),
    // An empty list is kept when the page said so: Delete on a workbook's
    // last custom style must reach the file it came from (XLS-A012).
    ...(customCellStyles.length > 0 || Array.isArray(source.customCellStyles)
      ? { customCellStyles }
      : {}),
    ...(chartTemplates.length > 0 ? { chartTemplates } : {}),
    ...(numberFormats ? { numberFormats } : {}),
    ...(source.calculationExceptDataTables === true
      ? { calculationExceptDataTables: true as const }
      : {}),
    ...(calculationIteration ? { calculationIteration } : {}),
    ...(source.precisionAsDisplayed === true ? { precisionAsDisplayed: true as const } : {}),
    ...(source.hideRuler === true ? { hideRuler: true as const } : {}),
    ...(customViews.length > 0 ? { customViews } : {}),
    spellCheck: source.spellCheck !== false,
  };
}

/** Display uses the actual package colours even when its theme is preserve-only. */
export function spreadsheetViewThemeAccents(options: SpreadsheetViewOptions): readonly string[] {
  if (options.nativeThemeDisposition === "preserve-only") {
    // Older imports lack a readable palette receipt. Their retained themeId
    // is not source authority: use the standard gallery fallback, never an
    // unrelated warm selection. A readable native receipt always wins.
    return options.nativeThemeAccents ?? paletteById("office").accents;
  }
  if (options.nativeThemeColors) {
    const { colors } = options.nativeThemeColors;
    return [
      colors.accent1,
      colors.accent2,
      colors.accent3,
      colors.accent4,
      colors.accent5,
      colors.accent6,
    ];
  }
  return paletteById(options.themeId).accents;
}

/** The on/off switches Excel keeps on each worksheet's view (sheetView
 *  @showGridLines, @showRowColHeaders, @showFormulas), each stored on the
 *  sheet as `key: true` when it differs from Excel's default. */
export type SpreadsheetSheetViewSwitch = "hideGridlines" | "hideHeadings" | "showFormulas";

/**
 * The sheets with one per-sheet view switch set (`key: true`) or cleared on
 * the chosen sheets, every other sheet reading as before; null when nothing
 * would change. `legacySet` is the switch stored for the whole workbook
 * (content level, before the switches were per sheet): it reads onto every
 * sheet, so a change writes it onto each one.
 */
export function sheetsWithViewSwitch<S extends object>(
  sheets: readonly S[],
  key: SpreadsheetSheetViewSwitch,
  set: boolean,
  chosen: (sheet: S, index: number) => boolean,
  legacySet = false,
): S[] | null {
  const reads = (sheet: S) => (sheet as Record<string, unknown>)[key] === true || legacySet;
  if (sheets.every((sheet, index) => !chosen(sheet, index) || reads(sheet) === set)) return null;
  return sheets.map((sheet, index) => {
    const on = chosen(sheet, index) ? set : reads(sheet);
    if (((sheet as Record<string, unknown>)[key] === true) === on) return sheet;
    const out = { ...sheet } as Record<string, unknown>;
    if (on) out[key] = true;
    else delete out[key];
    return out as S;
  });
}

/** A sheet's own view (sheetView@view): Normal unless it says Page Layout
 *  or Page Break Preview. */
export function sheetViewModeOf(sheet: { viewMode?: unknown } | null | undefined): SpreadsheetViewMode {
  const mode = sheet?.viewMode;
  return mode === "pageLayout" || mode === "pageBreakPreview" ? mode : "normal";
}

/**
 * Serialize view options for persistence at the content top level, emitting ONLY
 * the non-default fields so a default workbook stays byte-for-byte unchanged.
 * The caller spreads this into the content object (like `referenceStyle`).
 */
export function serializeViewOptions(
  options: SpreadsheetViewOptions,
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (options.hideFormulaBar) out.hideFormulaBar = true;
  if (options.customLists.length > 0) out.customLists = options.customLists;
  if (options.readingMode) out.readingMode = true;
  const focusCellColor = normalizeFocusCellColor(options.focusCellColor);
  if (focusCellColor) out.focusCellColor = focusCellColor;
  if (options.eyeCare) out.eyeCare = true;
  if (options.themeId && options.themeId !== "office") {
    out.themeId = options.themeId;
  }
  if (options.themeFontSetId && options.themeFontSetId !== "office") {
    out.themeFontSetId = options.themeFontSetId;
  }
  if (options.themeEffectId && options.themeEffectId !== "none") {
    out.themeEffectId = options.themeEffectId;
  }
  if (options.nativeThemeColors) {
    delete out.themeId;
    out.nativeThemeColors = options.nativeThemeColors;
  }
  if (options.nativeThemeFonts) {
    delete out.themeFontSetId;
    out.nativeThemeFonts = options.nativeThemeFonts;
  }
  if (options.nativeThemeEffects) {
    delete out.themeEffectId;
    out.nativeThemeEffects = options.nativeThemeEffects;
  }
  if (options.nativeThemeDisposition === "preserve-only") {
    out.nativeThemeDisposition = "preserve-only";
    if (options.nativeThemeAccents) out.nativeThemeAccents = [...options.nativeThemeAccents];
    if (options.nativeThemeReplaceable) out.nativeThemeReplaceable = true;
  }
  if (options.nativeTableStyles?.length) {
    out.nativeTableStyles = options.nativeTableStyles;
  }
  if (options.customCellStyles) {
    out.customCellStyles = options.customCellStyles;
  }
  if (options.calculationExceptDataTables) {
    out.calculationExceptDataTables = true;
  }
  if (options.calculationIteration) out.calculationIteration = { ...options.calculationIteration };
  if (options.precisionAsDisplayed) out.precisionAsDisplayed = true;
  if (options.hideRuler) out.hideRuler = true;
  if (options.customViews?.length) out.customViews = options.customViews;
  if (options.chartTemplates?.length) {
    out.chartTemplates = options.chartTemplates;
  }
  // Kept even when empty: an emptied list is the user's Delete, not silence.
  if (options.numberFormats) out.numberFormats = options.numberFormats;
  if (!options.spellCheck) out.spellCheck = false;
  return out;
}

/**
 * Decide whether a cell should render its FORMULA TEXT (Show Formulas mode) vs
 * its computed result. Returns true only when Show Formulas is ON and the cell's
 * stored value is a FORMULA (a non-empty string starting with "="). A plain
 * text/number value always shows as itself — Show Formulas only affects cells
 * that actually carry a formula, exactly like Excel (a literal "= sign" without
 * a formula, i.e. a value not starting with "=", is unaffected).
 *
 * The caller, when this returns true, displays the cell's stored value verbatim
 * (already converted to the active reference style by the existing
 * `displayFormula` path) instead of the evaluated result.
 */
export function shouldShowFormulaText(
  storedValue: string | undefined,
  showFormulas: boolean,
): boolean {
  return showFormulas && typeof storedValue === "string" && storedValue.startsWith("=");
}

/**
 * Whether a cell sits on the reading-mode crosshair: the active cell's row or
 * its column.
 *
 * The active cell itself is excluded — it already carries the selection ring,
 * and tinting it as well would read as a second, competing selection. A merged
 * anchor spans more than one line, so the whole span counts as "the active
 * row/column", which is what WPS highlights and what a user reading a merged
 * header expects.
 */
export function isOnReadingCrosshair(
  cell: { row: number; col: number },
  active: { row: number; col: number; rowSpan?: number; colSpan?: number },
): boolean {
  const rowSpan = Math.max(1, Math.floor(active.rowSpan ?? 1));
  const colSpan = Math.max(1, Math.floor(active.colSpan ?? 1));
  const onRow =
    cell.row >= active.row && cell.row < active.row + rowSpan;
  const onCol =
    cell.col >= active.col && cell.col < active.col + colSpan;
  if (onRow && onCol) return false;
  return onRow || onCol;
}
