/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
import { PX_PER_PT } from "../office-units";
import type { CSSProperties } from "react";
import { officePrimaryFontFamily } from "../../components/office-ribbon/officeTypographyCatalog";
import { officeFaceName } from "../office-fonts/face-names";
import { setPrintedCellScripts } from "./cell-text-effects";
import { excelCutFamilyEmboldens, excelCutFamilyOf, isExcelCutFamily, withoutLeadingExcelCutFamilies } from "../office-fonts/excel-cut-families";
import { sheetHangulEmFamily, stackPrimaryFamily } from "../office-fonts/hangul-em";
import { ensureOfficeFauxBoldProbe, officeFamilyHasNoBold } from "../office-fonts/faux-bold";
import { ensureExcelCutFamily, ensureOfficeFontFaces, loadOfficeFamilySpace, officeFontDeclared, officeFontSubstitutes, setOfficeEmboldenedRegularCuts } from "../office-fonts/substitutes";
import { columnBaseWidthToPixels, columnFileWidthToPixels } from "./grid-sizing";
import { isSpreadsheetFauxBoldStroke, setSpreadsheetFauxBoldPrinted, spreadsheetEngineIdealAdvances, spreadsheetEngineTextWidth, spreadsheetFauxBoldCss, spreadsheetTextEngineOn } from "./spreadsheet-text-engine";
import { withScriptFallbacks } from "../office-fonts/script-fallbacks";
import { OFFICE_COMPLEX_SCALE_PREFIX, officeComplexScaleStack } from "./office-complex-scales";
import { officeRegionalLettersStack } from "./office-regional-letters";
import { excelFontRecordOf, excelWindowsCellPitch, excelWindowsDrawnFaces, excelWindowsFaceKey, type ExcelFontRecordClass } from "../office-fonts/excel-windows-faces";
import OFFICE_DIGIT_WHOLE_EM from "./office-digit-whole-em.json";
import OFFICE_HINTED_ADVANCES from "./office-hinted-advances.json";
import { GENERIC_FAMILIES } from "../text-engine/face-index";
import { fontMeasureContext } from "./font-measure-context";

export { fontMeasureContext };
import { fontLoadReach } from "./font-load-reach";

/** SpreadsheetML widths are font-relative units, never CSS pixels.
 * Keep the source unit beside its legacy projection so merely opening a file
 * on a different display/font environment cannot rewrite the native package.
 * ISO 29500 CT_Col: round(width * MDW) on the target device's pixel grid.
 */
export interface NativeGridFont {
  family: string;
  sizePt: number;
  bold?: boolean;
  italic?: boolean;
  headingFamily?: string;
  bodyFamily?: string;
  themeRole?: "major" | "minor";
  /** The font record's class where Windows lacks the face it names: drawn and
   * measured in GDI's pick for it (excelFontRecordOf; reading
   * xlsxFontRecordClass). */
  record?: ExcelFontRecordClass;
}
export interface NativeEditProfile { dpi: 72 | 96; digitWidth: number }
export interface NativeColumnWidth { width: number; pixels: number; profile?: NativeEditProfile }
export interface NativeRowHeight { heightPt: number; pixels: number; customHeight: boolean }
export interface NativeRowMetrics { version: 1; default: NativeRowHeight; rows: Record<string, NativeRowHeight> }
export interface NativeGrid {
  version: 1;
  font: NativeGridFont;
  metricsFont?: NativeGridFont;
  columns: Record<string, NativeColumnWidth>;
  defaultWidth?: number;
  baseColumnWidth: number;
  defaultPixels: number;
  sourceThemeFontId?: string;
  defaultProfile?: NativeEditProfile;
  rowMetrics?: NativeRowMetrics;
}
const record = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const finite = (value: unknown, min: number, max: number): value is number =>
  typeof value === "number" && Number.isFinite(value) && value >= min && value <= max;
/** Excel's widest column, 255 characters, stored with the cell padding. */
const MAX_NATIVE_COLUMN_WIDTH = 256;
export function normalizeNativeGrid(value: unknown): NativeGrid | undefined {
  if (!record(value) || value.version !== 1 || !record(value.font) || !record(value.columns) ||
      typeof value.font.family !== "string" || !value.font.family.trim() || value.font.family.length > 200 ||
      /[\u0000-\u001f<>;{}"'\\]/u.test(value.font.family) || !finite(value.font.sizePt, 1, 409) ||
      !finite(value.baseColumnWidth, 0, 255) || !finite(value.defaultPixels, 1, 4096) ||
      (value.defaultWidth !== undefined && !finite(value.defaultWidth, 0, 255))) return undefined;
  const metricsFont = value.metricsFont === undefined ? undefined : normalizeNativeGrid({...value, font: value.metricsFont, metricsFont: undefined})?.font;
  if (value.metricsFont !== undefined && !metricsFont) return undefined;
  const profile = value.defaultProfile;
  const validProfile = (profile: unknown): profile is NativeEditProfile => record(profile) &&
    (profile.dpi === 72 || profile.dpi === 96) && finite(profile.digitWidth,1,512) && Number.isInteger(profile.digitWidth);
  if (profile !== undefined && !validProfile(profile)) return undefined;
  const entries = Object.entries(value.columns);
  if (entries.length > 16384) return undefined;
  const columns: Record<string, NativeColumnWidth> = {};
  for (const key of ["headingFamily", "bodyFamily"] as const) {
    const family = value.font[key];
    if (family !== undefined && (typeof family !== "string" || !family.trim() || family.length > 200 || /[\u0000-\u001f<>;{}"'\\]/u.test(family))) return undefined;
  }
  for (const [index, entry] of entries) {
    if (!/^(0|[1-9]\d*)$/.test(index) || Number(index) > 16383 || !record(entry) ||
        !finite(entry.width, 0, MAX_NATIVE_COLUMN_WIDTH) || !finite(entry.pixels, 1, 4096)) return undefined;
    if (entry.profile !== undefined && !validProfile(entry.profile)) return undefined;
    columns[index] = { width: entry.width, pixels: entry.pixels, ...(validProfile(entry.profile) ? { profile: entry.profile } : {}) };
  }
  const rowMetrics = normalizeNativeRowMetrics(value.rowMetrics);
  if (value.rowMetrics !== undefined && !rowMetrics) return undefined;
  return { version: 1, font: { family: value.font.family, sizePt: value.font.sizePt,
    ...(typeof value.font.headingFamily === "string" ? { headingFamily: value.font.headingFamily } : {}),
    ...(typeof value.font.bodyFamily === "string" ? { bodyFamily: value.font.bodyFamily } : {}),
    ...(value.font.themeRole === "major" || value.font.themeRole === "minor" ? { themeRole: value.font.themeRole } : {}),
    ...(value.font.bold === true ? { bold: true } : {}), ...(value.font.italic === true ? { italic: true } : {}),
    ...(excelFontRecordOf(value.font.record) ? { record: excelFontRecordOf(value.font.record) } : {}) },
    ...(metricsFont ? { metricsFont } : {}),
    columns, baseColumnWidth: value.baseColumnWidth, defaultPixels: value.defaultPixels,
    ...(typeof value.sourceThemeFontId === "string" && value.sourceThemeFontId.length <= 100 ? { sourceThemeFontId: value.sourceThemeFontId } : {}),
    ...(value.defaultWidth !== undefined ? { defaultWidth: value.defaultWidth } : {}),
    ...(validProfile(profile) ? { defaultProfile: profile } : {}),
    ...(rowMetrics ? {rowMetrics} : {}) };
}

export function normalizeNativeRowMetrics(value: unknown): NativeRowMetrics | undefined {
  const valid = (row: unknown): row is NativeRowHeight => record(row) && finite(row.heightPt,0.05,409.5) &&
    finite(row.pixels,1,4096) && typeof row.customHeight === "boolean";
  if (!record(value) || value.version !== 1 || !valid(value.default) || !record(value.rows) || Object.keys(value.rows).length > 200000) return undefined;
  const rows: Record<string,NativeRowHeight> = {};
  const copy = (row: NativeRowHeight) => ({heightPt:row.heightPt,pixels:row.pixels,customHeight:row.customHeight});
  for (const [key,row] of Object.entries(value.rows)) {
    if (!/^(0|[1-9]\d*)$/.test(key) || Number(key)>1048575 || !valid(row)) return undefined;
    rows[key]=copy(row);
  }
  return {version:1,default:copy(value.default),rows};
}

/** Pairing validates dimensions only; native write authority remains server-owned. */
export function nativeRowPixelPairMatches(metrics: NativeRowMetrics | undefined, index: string, pixels: unknown): boolean {
  const row = metrics?.rows[index];
  return Boolean(row && typeof pixels === "number" && pixels === row.pixels &&
    pixels === Math.max(1, Math.round(row.heightPt / .75)));
}

/** Pure view projection. Retain legacy saved pixels; only a matching pair may
 * use the original point height. Never feed this projection into autosave. */
export function resolveNativeRowHeights(metrics: NativeRowMetrics | undefined, heights: Record<string,number> | undefined, defaultHeight: number | undefined, fontRowHeight?: number, fontReplaced = false) {
  if (!metrics) return undefined;
  const resolved = {...heights};
  for (const [key,row] of Object.entries(metrics.rows)) if (heights?.[key] === row.pixels) resolved[key]=row.heightPt / .75;
  const pairedDefault = defaultHeight === metrics.default.pixels;
  const automatic = pairedDefault && !metrics.default.customHeight && finite(fontRowHeight, .05/.75, 4096);
  // A file saved by Excel at another DPI states its rows in that DPI's lines
  // (14.4 pt for a Calibri 11 line at 120 dpi): Windows Excel re-fits a row
  // without customHeight on open (lookwin p26b, p38), so such a row is left to
  // the automatic fit (lookwin probe p30, 182 rows: re-fitted 147 exact and 22
  // 1.5 pt off, where k stated lines standing k of the font's own gave 136 / 23;
  // R101-X17 row 42, 3 lines at 120 dpi, is 4 at 96, p37b). A file whose
  // default is Excel's own (96 dpi) keeps its rows. The API's
  // _native_print_sheet applies the same rule.
  const stated = metrics.default.heightPt / .75;
  // A default on Windows' whole pixels at 96 dpi is Excel's own line for the
  // file's own font, whatever the estimate of its face says (R464-X22's
  // Palatino Linotype 11 columns: 16.5 pt, 22 px, where the metrics gave 21);
  // a font a theme replaced is re-measured all the same.
  const otherDpi = fontReplaced || Math.abs(stated - Math.round(stated)) > .01;
  if (automatic && otherDpi && Math.abs(fontRowHeight - metrics.default.pixels) >= .5 && Math.abs(fontRowHeight - stated) >= .5) {
    for (const [key,row] of Object.entries(metrics.rows)) {
      if (row.customHeight || heights?.[key] !== row.pixels) continue;
      delete resolved[key];
    }
  }
  return {heights:resolved,defaultHeight:automatic ? clampNativeAutomaticRowHeight(fontRowHeight) : pairedDefault ? metrics.default.heightPt/.75 : defaultHeight};
}

/** Whether a default row the file marks custom (sheetFormatPr customHeight)
 * holds the rows that state no height of their own at it: only while no row
 * of the sheet states a custom height, or a height other than the default's.
 * Once one does, Windows Excel fits those rows to their content (lookwin
 * p97e-lw: a wrapped bold W92000004 in a row with no ht stands one line, two
 * once row 1, 3 or 5 states ht customHeight; R346-X27; p210b-lw / p210c-lw:
 * under a 15 pt custom default one row stating ht 30, 14, 38.25 or 12.75
 * without customHeight has every row fitted, where ht 15 leaves them 15 pt).
 * A hidden row does the same (`hidesRows`; R455-X33: six rows a filter hid,
 * no row stating a height, every row fitted). The API's writer
 * (_custom_default_row_fits) is the same. */
export function nativeDefaultHoldsRows(metrics: NativeRowMetrics | undefined, hidesRows = false): boolean {
  if (!metrics?.default.customHeight || hidesRows) return false;
  const stated = metrics.default.heightPt;
  return !Object.values(metrics.rows).some((row) => row.customHeight || Math.abs(row.heightPt - stated) > .005);
}

/** The height, in 96 dpi pixels at the sheet's zoom, of the grid Windows
 * Excel opens a workbook on (the truth machine's window): the API's
 * _EXCEL_OPEN_VIEW_GRID_PX. */
export const EXCEL_OPEN_VIEW_GRID_PX = 640;

/** Where a custom default holds its rows (nativeDefaultHoldsRows), how many
 * leading rows it holds: the ones Excel's opening view shows. Every recorded
 * row below them is fitted to its content after all, below the default too
 * (lookwin p34-lw: under a 15.75 pt custom default rows 1-31 stand 15.75 pt
 * and rows 32-80 are fitted; under 30 pt 16 rows; at zoom 50 59 rows;
 * R453-X62). undefined where the default holds no row. The API's
 * _custom_default_view_rows is the same. */
export function nativeDefaultViewRows(metrics: NativeRowMetrics | undefined, hidesRows = false, zoomScale?: unknown, freezeRows?: unknown, openingTopRow?: unknown): number | undefined {
  if (!metrics || !nativeDefaultHoldsRows(metrics, hidesRows)) return undefined;
  const height = metrics.default.heightPt;
  if (!(height > 0 && height <= 409.5)) return undefined;
  const zoom = typeof zoomScale === "number" && Number.isInteger(zoomScale) && zoomScale >= 10 && zoomScale <= 400 ? zoomScale / 100 : 1;
  const rowPx = Math.floor(Math.floor(height / .75 + .5) * zoom + .5);
  // Frozen rows stand above the scrolling pane's own view (p34-lw: ySplit 1
  // holds rows 1-32).
  const frozen = typeof freezeRows === "number" && Number.isInteger(freezeRows) && freezeRows > 0 ? freezeRows : 0;
  // A sheet that opens scrolled down (the page's openingTopRow, reading 955)
  // holds every row down to its view's last (p34-lw: A40 holds rows 1-70;
  // ySplit 5 with the bottom pane at A30, rows 1-60).
  const start = typeof openingTopRow === "number" && Number.isInteger(openingTopRow) && openingTopRow > frozen && openingTopRow < 1_048_576
    ? openingTopRow : frozen;
  return rowPx > 0 ? start + Math.ceil(EXCEL_OPEN_VIEW_GRID_PX / rowPx) : undefined;
}

/** The first row below the grid Windows Excel shows a sheet in when it opens
 * the workbook (EXCEL_OPEN_VIEW_GRID_PX at the sheet's zoom, below its frozen
 * rows and from its opening top row), measured by the rows' stated heights
 * (the file's ht, its last fit) or the default: a sheet at 100% fits those
 * rows at its own zoom where the rest take the zoom of the zoomed-out sheet
 * the workbook opens on (lookwin p258/p260-lw). The API's
 * _open_view_rows_by_stated. */
export function nativeOpenViewRowsByStated(metrics: NativeRowMetrics | undefined, zoomScale?: unknown, freezeRows?: unknown, openingTopRow?: unknown): number {
  const fallback = metrics?.default.heightPt;
  const defaultPt = typeof fallback === "number" && fallback > 0 && fallback <= 409.5 ? fallback : 15;
  const zoom = typeof zoomScale === "number" && Number.isInteger(zoomScale) && zoomScale >= 10 && zoomScale <= 400 ? zoomScale / 100 : 1;
  const frozen = typeof freezeRows === "number" && Number.isInteger(freezeRows) && freezeRows > 0 && freezeRows < 1_048_576 ? freezeRows : 0;
  let row = typeof openingTopRow === "number" && Number.isInteger(openingTopRow) && openingTopRow > frozen && openingTopRow < 1_048_576
    ? openingTopRow : frozen;
  let reach = 0;
  while (reach < EXCEL_OPEN_VIEW_GRID_PX && row < 1_048_576) {
    const stated = metrics?.rows[String(row)]?.heightPt;
    const height = typeof stated === "number" && stated >= 0 && stated <= 409.5 ? stated : defaultPt;
    reach += Math.floor(Math.floor(height / .75 + .5) * zoom + .5);
    row += 1;
  }
  return row;
}

/** Excel's row limit applies to computed heights as well as authored heights. */
export function clampNativeAutomaticRowHeight(height: number): number {
  return Math.min(409.5 / .75, height);
}

/** Synchronize only explicit canonical dimension changes. Structural writers
 * remap the original pairs first; display/font measurement never calls this. */
export function syncNativeRowMetrics(grid: NativeGrid, heights: Record<string,number> | undefined, defaultHeight: number | undefined): NativeGrid {
  const source=grid.rowMetrics;
  if (!source) return grid;
  const rows:Record<string,NativeRowHeight>={};
  for (const [key,pixels] of Object.entries(heights??{})) {
    rows[key]=source.rows[key]?.pixels===pixels ? source.rows[key] : {heightPt:pixels*.75,pixels,customHeight:true};
  }
  const rowDefault=defaultHeight===undefined || defaultHeight===source.default.pixels ? source.default : {heightPt:defaultHeight*.75,pixels:defaultHeight,customHeight:true};
  return {...grid,rowMetrics:{version:1,default:rowDefault,rows}};
}

export function editNativeRowSizes(grid: NativeGrid, sizes: readonly {index:number;size:number}[]): NativeGrid {
  if (!grid.rowMetrics) return grid;
  const rows={...grid.rowMetrics.rows};
  for (const {index,size} of sizes) {
    const bounded=clampNativeRowHeight(size,20);
    rows[index]={heightPt:bounded*.75,pixels:Math.round(bounded),customHeight:true};
  }
  return {...grid,rowMetrics:{...grid.rowMetrics,rows}};
}

export function clampNativeRowHeight(value: unknown, fallback: number): number {
  const size=typeof value==="number" && Number.isFinite(value) ? value : Number.isFinite(fallback) ? fallback : 20;
  return Math.max(18,Math.min(512,size));
}

export function nativeRowCssClamp(value: unknown, fallback: number): number {
  return finite(value,.05/.75,4096) ? value : finite(fallback,.05/.75,4096) ? fallback : 20;
}
export function nativeWidthToPixels(width: number, digitWidth: number, dpi: number): number {
  return Math.max(1, columnFileWidthToPixels(width, digitWidth)) * 96 / dpi;
}
/** CT_Col reserves two device pixels of text padding on each side. Keep
 * that physical inset on the same device grid as native glyph advances. */
export function nativeCellPaddingPx(dpi: number): number {
  return 2 * 96 / dpi;
}
export function pixelsToNativeWidth(pixels: number, digitWidth: number, dpi: number): number {
  // Choose the smallest 1/256 native unit that roundtrips to this device pixel.
  const device = Math.max(1, Math.round(pixels * dpi / 96));
  return Math.max(0, Math.min(255, Math.ceil((device * 256 / digitWidth) - Math.floor(128 / digitWidth)) / 256));
}
export function nativeDefaultWidth(grid: NativeGrid, digitWidth: number, dpi: number, sourceDigitWidth?: number) {
  if (grid.defaultWidth !== undefined) return nativeWidthToPixels(grid.defaultWidth, digitWidth, dpi);
  // A `<col>` over every column, which no sheetFormatPr states, is the
  // sheet's default (the reader's page_projection.materialize_columns): its
  // pixels are not the base width's in the file's own digit. Every column it
  // covers keeps its width, as Excel draws and prints them (lookwin R211-X05:
  // `<col min="1" max="16384" width="8.7265625">`, 61 px, drawn 64 past the
  // used range and the text box over them 24 pt wider than Excel's).
  if (
    sourceDigitWidth !== undefined &&
    sourceDigitWidth > 0 &&
    columnBaseWidthToPixels(grid.baseColumnWidth, sourceDigitWidth, dpi) * 96 / dpi !== grid.defaultPixels
  ) {
    return nativeWidthToPixels(pixelsToNativeWidth(grid.defaultPixels, sourceDigitWidth, dpi), digitWidth, dpi);
  }
  return columnBaseWidthToPixels(grid.baseColumnWidth, digitWidth, dpi) * 96 / dpi;
}
export function nativeColumnCssClamp(value: unknown, fallback: number): number {
  return finite(value, 0, 4096) ? Math.max(1, value) : fallback;
}

export function editNativeColumns(grid: NativeGrid, sizes: readonly { index: number; size: number }[], digitWidth: number, dpi: number): NativeGrid {
  const columns = { ...grid.columns };
  for (const { index, size } of sizes) {
    const width = pixelsToNativeWidth(size, digitWidth, dpi);
    columns[index] = { width, pixels: nativeWidthToPixels(width,digitWidth,dpi), profile: { dpi: dpi as 72 | 96, digitWidth } };
  }
  return { ...grid, columns };
}

const faceAvailability = new Map<string, boolean>();
/** Bumped whenever a font finishes loading (the face and advance caches are
 * dropped then): text drawn before it may be measured anew. */
let typographyGeneration = 0;
export function nativeTypographyGeneration(): number {
  return typographyGeneration;
}
const typographyListeners = new Set<() => void>();
let fontLoadsWatched = false;
function watchFontLoads(): void {
  if (fontLoadsWatched || typeof document === "undefined" || !document.fonts) return;
  fontLoadsWatched = true;
  document.fonts.addEventListener("loadingdone", (event) => {
    // Only the families the load finished faces of can have changed; one
    
    // fallback slice moved the generation, and every measure keyed by it).
    const reach = fontLoadReach(event);
    let changed = !reach;
    if (!reach) faceAvailability.clear();
    else {
      for (const [family, available] of [...faceAvailability]) {
        if (!reach.families.has(family.replace(/^["']|["']$/g, "").trim().toLowerCase())) continue;
        faceAvailability.delete(family);
        if (!available) changed = true;
      }
    }
    if (!changed) return;
    typographyGeneration += 1;
    for (const listener of [...typographyListeners]) listener();
  });
}




export function subscribeNativeTypography(listener: () => void): () => void {
  watchFontLoads();
  typographyListeners.add(listener);
  return () => {
    typographyListeners.delete(listener);
  };
}
/** A family a text waits for (drawn meanwhile as it can): its arrival moves
 * the generation, as a face read as unavailable does (office-arabic-stretch). */
export function awaitSpreadsheetFamily(family: string): void {
  if (typeof document === "undefined") return;
  watchFontLoads();
  if (faceAvailability.size > 512) faceAvailability.clear();
  faceAvailability.set(family, false);
}
let availabilityContext: CanvasRenderingContext2D | null | undefined;
export function spreadsheetFontAvailable(family: string): boolean | undefined {
  if (typeof document === "undefined") return undefined;
  if (availabilityContext === undefined) {
    availabilityContext = document.createElement("canvas").getContext("2d");
    watchFontLoads();
  }
  if (!availabilityContext) return undefined;
  if (faceAvailability.has(family)) return faceAvailability.get(family);
  // A face the product serves itself (styles/office-fonts*.css alias Office's
  // families to open faces behind local()) is available wherever the page
  // is, even before its file has arrived: the width probe below would read
  // the interim fallback and pick the substitution stack for a font that is
  // on its way. Ask for it now so the swap lands with the first paint it
  // can. A family whose stand-in the stack names itself counts the same.
  if (officeFontDeclared(family)) {
    void document.fonts.load(`12px "${family}"`).catch(() => undefined);
    faceAvailability.set(family, true);
    return true;
  }
  if (officeFontSubstitutes(family).length) {
    faceAvailability.set(family, true);
    return true;
  }
  // FontFaceSet.check returns true for unknown local families. Independent
  // fallback probes distinguish a real selected face from silent substitution.
  const widths = ["serif", "sans-serif", "monospace"].map(fallback => {
    availabilityContext!.font = `100px "${family}", ${fallback}`;
    return availabilityContext!.measureText("0123456789MWil").width;
  });
  const available = Math.max(...widths) - Math.min(...widths) < 0.01;
  if (faceAvailability.size > 512) faceAvailability.clear();
  faceAvailability.set(family, available);
  return available;
}

const LATIN_SERIF_FACE = /^(times( new roman)?|georgia|cambria|garamond|book antiqua|palatino( linotype)?|century( schoolbook)?|constantia|bookman old style|baskerville( old face)?|perpetua|goudy old style|bodoni mt|didot|hoefler text|charter|iowan old style)$/i;
const LATIN_SERIF_STAND_IN = /^(tinos|caladea|gelasio|literata|bodoni moda|rokkitt|cormorant garamond|eb garamond|tex gyre pagella|cardo|alegreya|liberation serif|crimson (text|pro)|libre baskerville)$/i;

/** A Latin face drawn with serifs, by name or by the open face standing in for it. */
function latinSerifFace(face: string, standIns: readonly string[]): boolean {
  return LATIN_SERIF_FACE.test(face) || (standIns.length > 0 && LATIN_SERIF_STAND_IN.test(standIns[0]));
}

/** The same declared face/fallback stack is used by canvas, DOM and AutoFit.
 * Keep Office Korean coverage after the established Han/Kana fallbacks, so
 * missing Hangul glyphs do not fall through to a different system font.
 */
// Every cell of every render asks for its stack, and building one walks
// substitutes, script fallbacks and the served faces (in Page Layout, about
// 0.6 s of each scroll's re-render). A stack depends only on the family, the
// traits, the text engine switch and what the browser has loaded, which
// advances `typographyGeneration`, so it is kept per generation.
const fontStackCache = new Map<string, string>();
let fontStackCacheGeneration = -1;

export function spreadsheetFontStack(
  family: string | undefined,
  traits: Pick<NativeGridFont, "bold" | "italic"> & { readonly record?: unknown; readonly fontRecord?: unknown } = {},
): string {
  // A face Windows lacks whose record states its class (a grid font's
  // `record`, a cell's `fontRecord`) is drawn in GDI's pick for that class:
  // Carlito stating nothing in Arial, charset 134 in SimSun.
  // A stack built here and handed back: its face, not the cut family in front of it.
  if (family) family = withoutLeadingExcelCutFamilies(family);
  const record = excelFontRecordOf(traits.fontRecord ?? traits.record);
  if (record) {
    const drawn = excelWindowsDrawnFaces(family?.trim(), record);
    if (drawn) family = drawn.latin;
  }
  const engine = spreadsheetTextEngineOn();
  if (fontStackCacheGeneration !== typographyGeneration) {
    fontStackCache.clear();
    fontStackCacheGeneration = typographyGeneration;
  }
  const mac = typeof navigator !== "undefined" && /Mac/i.test(navigator.platform);
  const key = `${engine ? 1 : 0}${mac ? 1 : 0}${traits.bold ? 1 : 0}${traits.italic ? 1 : 0}|${family ?? ""}`;
  const cached = fontStackCache.get(key);
  if (cached !== undefined) return cached;
  // Traits stay in the key: a cascaded bold asks again (resolveSpreadsheetFontStyle).
  const stack = buildSpreadsheetFontStack(family);
  if (typeof document !== "undefined") {
    if (fontStackCache.size >= 4096) fontStackCache.clear();
    fontStackCache.set(key, stack);
  }
  return stack;
}

function buildSpreadsheetFontStack(family: string | undefined): string {
  // A cell stores one face, but the shared font picker hands a CSS stack
  // ("Arial, Helvetica, sans-serif"): its first family is the face. Read
  // whole, the stack was one unknown name drawn in the serif fallback (user
  // audit S-006).
  const named = family?.split(",")[0]?.replace(/["'\\;{}()<>]/g, "").trim().slice(0, 200);
  // A name Windows Excel lacks or aliases is drawn in its pick whatever this
  // browser has: Carlito, Helvetica Neue and "Calibri (Body)" in Arial,
  // Helvetica in Arial, Times in Times New Roman (lookwin faces probe; the
  // name as the file states it, else the stack's first face).
  const windows = excelWindowsDrawnFaces(family?.trim()) ?? excelWindowsDrawnFaces(named);
  // A face named by its full or PostScript name ("Lucida Sans Regular") is
  // that face of its family, as Excel for Mac opens it (X-028).
  const face = windows
    ? windows.latin
    : named && spreadsheetFontAvailable(named) === false ? officeFaceName(named)?.family ?? named : named;
  // The face stands first in the stack: its slice holding the space is the
  // line's first available font once loaded (loadOfficeFamilySpace).
  loadOfficeFamilySpace(face);
  if (!windows && face && spreadsheetFontAvailable(face) === false) {
    // A face nothing here draws: Windows Excel draws a font record it lacks
    // in Arial (one that states no class, as most do), Han and kana in
    // Microsoft YaHei and Hangul in Malgun Gothic, bold and italic alike
    // (lookwin faces probe, 542 faces). Keep the requested family first so a
    // face that arrives later still wins.
    ensureOfficeFontFaces("Arial");
    const arial = officeFontSubstitutes("Arial").map((name) => `"${name}", `).join("");
    return servedStack(`"${face}", "Arial", ${arial}${officeComplexScaleStack(face)}${officeRegionalLettersStack()}"Microsoft YaHei", "Malgun Gothic", "Mangal", "Tahoma", sans-serif`);
  }
  ensureOfficeFontFaces(face);
  // A face Windows Excel prints lighter than the catalog's cut of its
  // stand-in (游ゴシック Medium: Noto Sans JP's 500, half as heavy again as
  // Excel's ink) is drawn by the sheet's own family of it, named in front of
  // the face (lib/office-fonts/excel-cut-families.ts): the measured cut at
  // the same advances. It stands first, so its space is the line's first
  // available font.
  const cut = excelCutFamilyOf(face);
  if (cut) {
    ensureExcelCutFamily(cut.face);
    loadOfficeFamilySpace(cut.alias);
  }
  // A Korean face's Hangul at its full em, as Windows Excel sets it in
  // every Korean face (the stand-in's is 0.92 em): the sheet's own family
  // of its clone, declared with the face (hangul-em.ts sheetHangulEmFamily).
  // It holds Hangul alone; every other character falls through to the face.
  const hangul = cut ? null : sheetHangulEmFamily(face);
  const standIns = face ? officeFontSubstitutes(face) : [];
  const substitutes = standIns.map((name) => `"${name}", `).join("");
  // Han text in a Latin face has no glyphs of its own and takes the system's
  // East Asian fallback, which keeps the face's class: Excel draws the Han in
  // a Times New Roman cell as Song (serif) and in Calibri or Arial as a sans
  // (measured, Excel for Mac). Only a serif Latin face gets the serif Han.
  const hanSerif = face && latinSerifFace(face, standIns) ? (ensureOfficeFontFaces("SimSun"), '"SimSun", "Songti SC", ') : "";
  // A serif face's Cyrillic or Greek its stand-in lacks (Caladea, Cambria's,
  // has none) is drawn by a proportional serif with those scripts, not by
  // the Han face after it, whose Cyrillic is wide and spaced (PH057's
  // Cambria "Ключевые характеристики", X-246).
  const scriptSerif =
    hanSerif && !/^times new roman$/i.test(face ?? "")
      ? (ensureOfficeFontFaces("Times New Roman"), '"Times New Roman", ')
      : "";
  // Arabic, Hebrew and Bengali at the widths Windows Excel sets them in
  // this face, right behind the face and ahead of its stand-ins
  // (office-complex-scales.ts), as the API's writer prints them: a stand-in
  // holding the script draws it at its own width (Tinos' Hebrew for Times
  // New Roman, 10% wider than Excel's), and what the stand-ins lack would
  // fall to Wine Tahoma's Arabic, behind every face here, or to the script
  // fallback's Noto at its own size. Each family covers its script alone.
  const complex = officeComplexScaleStack(face);
  // Regional indicators as Excel's letters (office-regional-letters.ts),
  // ahead of every stand-in and fallback that would draw a flag or a box.
  const regional = officeRegionalLettersStack();
  return servedStack(`${face ? `${cut ? `"${cut.alias}", ` : hangul ? `"${hangul}", ` : ""}"${face}", ${complex}${regional}${substitutes}${scriptSerif}${hanSerif}` : ""}"Calibri","Microsoft YaHei", "Malgun Gothic", "Mangal", "Tahoma", "PingFang SC", sans-serif`);
}

/** The system faces a stack's Han and kana fall to when its own faces lack them. */
const HAN_FALLBACK_FACES = new Set(["simsun", "songti sc", "microsoft yahei", "malgun gothic", "pingfang sc"]);

/**
 * The stack a sheet SHAPE's words in `family` are drawn in. A Latin face
 * whose run leaves its East Asian face to the theme (`a:ea typeface=""`) has
 * its Han and kana drawn in the theme's face for the run's language, which in
 * a workbook whose Normal font is East Asian is that font (Windows Excel
 * prints R361-X03's Calibri text boxes' 「６９２，９６８百万円」 in
 * ＭＳ Ｐゴシック, proportional, not a full-em fallback; the API's print route
 * `_sheet_drawing_face_resolver`). `eastAsianFace` is that Normal font (none
 * for a Latin Normal font): its faces stand behind the Latin face's own and
 * ahead of the system's Han fallback, so the Latin letters keep their face.
 */
export function spreadsheetShapeFontStack(family: string | undefined, eastAsianFace?: string | null): string {
  const stack = spreadsheetFontStack(family);
  const named = family ? withoutLeadingExcelCutFamilies(family).split(",")[0]?.replace(/["']/g, "").trim() : undefined;
  if (!eastAsianFace || !named || named === eastAsianFace) return stack;
  const parts = stack.split(",").map((part) => part.trim()).filter(Boolean);
  const bare = (part: string) => part.replace(/^["']|["']$/g, "").toLowerCase();
  const at = parts.findIndex((part, index) => index > 0 && HAN_FALLBACK_FACES.has(bare(part)));
  if (at < 0) return stack;
  const head = new Set(parts.slice(0, at).map(bare));
  const eastAsian: string[] = [];
  for (const part of spreadsheetFontStack(eastAsianFace).split(",").map((piece) => piece.trim())) {
    const name = bare(part);
    // The East Asian face's own faces, up to the shared fallback tail.
    if (!name || HAN_FALLBACK_FACES.has(name) || GENERIC_FAMILIES.has(name)) break;
    if (head.has(name) || name.startsWith(OFFICE_COMPLEX_SCALE_PREFIX.toLowerCase()) || name === "calibri") continue;
    eastAsian.push(part);
    head.add(name);
  }
  if (!eastAsian.length) return stack;
  return [...parts.slice(0, at), ...eastAsian, ...parts.slice(at)].join(", ");
}

/**
 * With the text engine on, a cell's stack ends in faces the page serves
 * rather than in whatever the visitor's system has: every Office family in
 * it is declared, the system's PingFang gives way to the served script
 * fallback family, and only then the generic family. Text the engine
 * measures is drawn in exactly the faces it measured (lib/text-engine).
 */
function servedStack(stack: string): string {
  if (!spreadsheetTextEngineOn()) return stack;
  const parts = stack.split(",").map((part) => part.trim()).filter((part) => part && !/^["']?PingFang SC["']?$/i.test(part));
  for (const part of parts) {
    const name = part.replace(/^["']|["']$/g, "");
    if (!GENERIC_FAMILIES.has(name.toLowerCase()) && !name.startsWith(OFFICE_COMPLEX_SCALE_PREFIX) && !isExcelCutFamily(name)) ensureOfficeFontFaces(name);
  }
  return withScriptFallbacks(parts.join(", "));
}

/** Characters every sheet draws: its numbers, their signs and separators. */
const WARM_TEXT = "0123456789 .,-+%$()#E/:";











export function warmSpreadsheetFamilyMeasures(family: string): boolean {
  if (!spreadsheetTextEngineOn() || !officeFontDeclared(family)) return false;
  for (const bold of [false, true]) {
    spreadsheetEngineTextWidth(WARM_TEXT, `normal ${bold ? 700 : 400} 15px ${spreadsheetFontStack(family, { bold })}`);
  }
  return true;
}

/** Resolve after the style cascade, including conditional/table and rich-run
 * traits. Measurement and paint must select the same actual fallback face. */
export function resolveSpreadsheetFontStyle(style: CSSProperties): CSSProperties {
  return withOfficeFauxBold(resolveSpreadsheetFontFamily(style));
}

/** Bold in an Office family with no bold face, decided on the cascaded
 * weight (a table style's or a rule's bold too, and a rule that turns a bold
 * cell regular): the faux-bold stroke where the browser synthesizes none
 * (lib/office-fonts/faux-bold; S11_totalsRowFunction's table-style totals
 * row, X-099). */
function withOfficeFauxBold(style: CSSProperties): CSSProperties {
  const bold = style.fontWeight === "bold" || Number(style.fontWeight) >= 600;
  const family = typeof style.fontFamily === "string" ? style.fontFamily : undefined;
  if (bold && (officeFamilyHasNoBold(family) || excelCutFamilyEmboldens(family))) {
    const faux = spreadsheetFauxBoldCss();
    if (style.WebkitTextStrokeWidth === faux.WebkitTextStrokeWidth && style.fontSynthesis === faux.fontSynthesis) return style;
    ensureOfficeFauxBoldProbe();
    return { ...style, ...faux };
  }
  if (!isSpreadsheetFauxBoldStroke(style.WebkitTextStrokeWidth)) return style;
  const rest = { ...style };
  delete rest.WebkitTextStrokeWidth;
  delete rest.fontSynthesis;
  return rest;
}

function resolveSpreadsheetFontFamily(style: CSSProperties): CSSProperties {
  if (typeof style.fontFamily !== "string") return style;
  const family = officePrimaryFontFamily(style.fontFamily);
  if (!family || spreadsheetFontAvailable(family) !== false || !/Mac/i.test(navigator.platform)) return style;
  const fontFamily = spreadsheetFontStack(family, {
    bold: style.fontWeight === "bold" || Number(style.fontWeight) >= 600,
    italic: style.fontStyle === "italic" || style.fontStyle?.startsWith("oblique") === true,
  });
  return fontFamily === style.fontFamily ? style : { ...style, fontFamily };
}

/** Native Mac substitution follows the leading script of a source font run.
 * Paired Han/Latin controls with two absent families retain italic in the
 * source, but paint Han-led italic runs with regular fallback faces. Resolve
 * before wrapping and carry the original leading text onto every fragment.
 * Available/bound fonts and other native device profiles keep their traits.
 */
export function resolveSpreadsheetTextStyle(style: CSSProperties, text: string, dpi?: number): CSSProperties {
  let resolved = resolveSpreadsheetFontStyle(style);
  // Mac screen-copy vectors quantize the drawing size before positioning
  // glyphs, for both available faces and substitutes. Keep this a disposable
  // CSS projection: imported point sizes and their native pairs stay exact.
  if (dpi === 72 && typeof navigator !== "undefined" && /Mac/i.test(navigator.platform) &&
      typeof resolved.fontSize === "number" && Number.isFinite(resolved.fontSize)) {
    const fontSize = Math.round(resolved.fontSize * .75) / .75;
    if (fontSize > 0 && fontSize !== resolved.fontSize) resolved = { ...resolved, fontSize };
  }
  if (dpi !== 72 || typeof navigator === "undefined" || !/Mac/i.test(navigator.platform) ||
      style.fontStyle !== "italic" || style.fontWeight === "bold" || Number(style.fontWeight) >= 600 ||
      !/^\p{Script=Han}/u.test(text) || typeof style.fontFamily !== "string") return resolved;
  const family = officePrimaryFontFamily(style.fontFamily);
  if (!family || spreadsheetFontAvailable(family) !== false) return resolved;
  return { ...resolved, fontFamily: spreadsheetFontStack(family), fontStyle: "normal" };
}

let textLayoutQuantum: number | undefined;
/** Detect the layout engine's fractional CSS grid instead of assuming the
 * same subpixel denominator across browsers. Independent of display DPR. */
export function nativeTextLayoutQuantum(): number {
  if (textLayoutQuantum !== undefined) return textLayoutQuantum;
  if (typeof document === "undefined" || !document.body) return 0;
  const probe = document.createElement("span");
  probe.style.cssText = "position:fixed;visibility:hidden;pointer-events:none;display:block;box-sizing:content-box;border:0;padding:0;margin:0;height:0;contain:strict;";
  document.body.appendChild(probe);
  let smallest = 0;
  let cssScale = 1;
  for (let power = 0; power <= 16; power += 1) {
    probe.style.width = `${1 / 2 ** power}px`;
    const width = probe.getBoundingClientRect().width;
    if (power === 0 && width > 0) cssScale = width;
    if (width > 0) smallest = width / cssScale;
    else break;
  }
  probe.remove();
  textLayoutQuantum = smallest;
  return smallest;
}

let textContext: CanvasRenderingContext2D | null | undefined;
const advanceCache = new Map<string, number>();
const idealAdvanceCache = new Map<string, number>();
/** Whole answers by text, font, grid, scale and offset: a scrolled band asks
 * again for what the rows it reuses were given (the same values down a
 * column, a row rendered again for its interaction). Same lifetime as the
 * per-glyph caches. */
const advancesCache = new Map<string, ReadonlyArray<{ text: string; spacing: number; advance?: number }>>();
let graphemes: Intl.Segmenter | undefined;
/** Excel device-pixel advances, independently measured for each grapheme.
 * Joining scripts retain browser shaping as a run; no code-unit splitting.
 */
/** A font load drops the advances it can change: those of a font and text
 * it reaches (font-load-reach), every one when it names no faces. True when
 * any went. */
export function dropAdvancesTouched(event: Event): boolean {
  const reach = fontLoadReach(event);
  if (!reach) {
    advanceCache.clear(); idealAdvanceCache.clear(); advancesCache.clear();
    return true;
  }
  let dropped = false;
  for (const cache of [advanceCache, idealAdvanceCache]) {
    for (const key of [...cache.keys()]) {
      const split = key.indexOf(":");
      if (split < 0 || reach.touchesFont(key.slice(0, split), key.slice(split + 1))) {
        cache.delete(key);
        dropped = true;
      }
    }
  }
  for (const key of [...advancesCache.keys()]) {
    const [text, font] = key.split("\u0000");
    if (font === undefined || reach.touchesFont(font, text)) {
      advancesCache.delete(key);
      dropped = true;
    }
  }
  return dropped;
}

/**
 * One indent level as Windows Excel prints it, CSS px of the paper: three of
 * the Normal font's spaces, each a whole printer dot at 600 dpi, at every
 * page scale -- the indent does not shrink with the page (lookwin p105-lw:
 * 8.27 pt a level in an Arial 10 workbook at 100, 50 and 25 % and Fit To;
 * p87-lw: Calibri 11 63 dots, Arial 10 69). The space's design width from
 * the face's file where the text engine has it, else the browser's.
 * Undefined without a face or a measure.
 */
export function officePrintedIndentStepPx(family: string | undefined, sizePt: number | undefined): number | undefined {
  if (typeof document === "undefined" || !family || typeof sizePt !== "number" || !(sizePt > 0)) return undefined;
  const font = `normal 400 1000px ${spreadsheetFontStack(family)}`;
  const em = spreadsheetEngineIdealAdvances([" "], font)?.[0] ?? fontMeasureContext(font, "none")?.measureText(" ").width;
  if (typeof em !== "number" || !(em > 0)) return undefined;
  const dots = Math.round(em / 1000 * sizePt * 600 / 72);
  return dots > 0 ? 3 * dots * 96 / 600 : undefined;
}

/**
 * The printer resolution the editor's print page paints text at, or 0 on
 * screen. Windows Excel prints a run's glyphs at the printer's whole pixels
 * per em and each advance in whole printer dots (600 dpi: Calibri 11 at 92
 * px per em, its digit 47 dots, office-print-metrics; lookwin p100c / p102g:
 * a long note runs 0.5% past its design width), not at the screen's whole
 * pixels (Calibri 11's digit 7 px: 5.25 pt where the paper shows 5.64). Lines
 * break where the screen breaks them ("measure"), but a printed wrapped
 * paragraph, which the printer re-wraps at these advances (the engine's
 * layout under NATIVE_PRINTED_ROOM_MARK, R219-X03). The API's sheet PDF writer is the same
 * (notebooks.py FaceMeasure.printed).
 */
let printerTextDpi = 0;
/** The printer's resolution text is painted at (0 off the print page). */
export function nativeTextPrinterDpi(): number {
  return printerTextDpi;
}
export function setNativeTextPrinterDpi(dpi: number): void {
  const next = Number.isFinite(dpi) && dpi > 96 ? dpi : 0;
  if (next === printerTextDpi) return;
  printerTextDpi = next;
  advancesCache.clear();
  // Bold in a face without a bold cut prints as Excel prints it too.
  setSpreadsheetFauxBoldPrinted(next > 0);
  // ... in its stand-in's heavier cut where Excel's is heavier than it.
  setOfficeEmboldenedRegularCuts(next > 0);
  // Superscripts and subscripts at Excel's printed size and rise.
  setPrintedCellScripts(next > 0);
}











const designedTextOwners = new Map<object, ReadonlySet<string>>();
export function setNativeTextDesignedLines(owner: object, lines: ReadonlySet<string> | null): void {
  if (lines && lines.size) designedTextOwners.set(owner, lines);
  else designedTextOwners.delete(owner);
}
function designedTextLine(text: string): boolean {
  for (const lines of designedTextOwners.values()) if (lines.has(text)) return true;
  return false;
}

/**
 * The print windows whose sheet prints on a landscape page, by owner (as
 * setNativeTextDesignedLines keeps its lines). File > Print through Microsoft
 * Print to PDF turns a landscape page's text itself, and a pure-ASCII string
 * of a GDI East Asian face (digits and Latin half an em: ＭＳ ゴシック, ＭＳ 明朝,
 * ＭＳ Ｐゴシック, ＭＳ Ｐ明朝, MS UI Gothic, 宋体, 新宋体) at 51 printer dots an em
 * or more then stands its half-dot advances rounded down -- 37 dots a digit
 * at 9 pt, where a portrait page, a smaller em and every string holding a
 * full-width character stand them 38 -- while Excel still places the text by
 * the rounded-up width, so a right-aligned number ends that many dots short
 * of its inset (lookwin p202-lw to p205-lw; R70-X09). The API's writer is the
 * same (notebooks.py FaceMeasure.printed `set_short`). The page's windows
 * share one text cache: the rule is on while every registered window is
 * landscape (a workbook of both orientations keeps the rounded-up advances).
 */
const landscapePrintOwners = new Map<object, boolean>();
export function setNativeTextLandscapePrint(owner: object, landscape: boolean | null): void {
  const before = nativeTextLandscapePrint();
  if (landscape === null) landscapePrintOwners.delete(owner);
  else landscapePrintOwners.set(owner, landscape);
  if (before !== nativeTextLandscapePrint()) advancesCache.clear();
}
function nativeTextLandscapePrint(): boolean {
  if (!landscapePrintOwners.size) return false;
  for (const landscape of landscapePrintOwners.values()) if (!landscape) return false;
  return true;
}
/** A GDI East Asian face whose digits and Latin are half an em (office-digit-whole-em.json; Gill Sans Nova, its one
 * Western face, prints rounded up like Times New Roman's half-em digits). */
function halfEmAsciiFace(family: string | null | undefined): boolean {
  const own = (family ?? "").trim().replace(/^["']|["']$/g, "").trim();
  if (!own) return false;
  const faces = OFFICE_DIGIT_WHOLE_EM as readonly string[];
  return [own.toLowerCase(), excelWindowsFaceKey(own)].some((name) => name !== "gill sans nova" && faces.includes(name));
}

/**
 * office-hinted-advances.json (the API's office_hinted_advances.json is the
 * same file): the printer dots Windows Excel's print sets an ASCII character
 * of a face wider or narrower than its design advance rounded to a whole dot,
 * per whole printer ppem (33-134; to 300 for the faces measured there) --
 * GDI's hinted advances, which the old TrueType faces do not keep linear
 * (Times New Roman's y prints 1-4 dots short at 90 of those sizes, Arial's W
 * 1-6 over at 100; lookwin p206-lw / p207-lw). The table of `family` (its
 * own name, or the face Windows Excel draws it in) regular, bold, italic or
 * bold italic: ppem -> character -> dots; null for a face or a weight not
 * measured.
 */
const hintedAdvanceTables = new Map<string, ReadonlyMap<number, ReadonlyMap<string, number>> | null>();
export function officeHintedAdvanceDeltas(family: string | null | undefined, bold: boolean, italic: boolean): ReadonlyMap<number, ReadonlyMap<string, number>> | null {
  const own = (family ?? "").trim().replace(/^["']|["']$/g, "").trim();
  if (!own) return null;
  const style = `${bold ? "B" : ""}${italic ? "I" : ""}` || "R";
  const faces = (OFFICE_HINTED_ADVANCES as unknown as { faces: Record<string, Record<string, Record<string, Record<string, string>>> | undefined> }).faces;
  for (const key of [own.toLowerCase(), excelWindowsFaceKey(own)]) {
    const cached = `${key}\u0000${style}`;
    let table = hintedAdvanceTables.get(cached);
    if (table === undefined) {
      const stated = faces[key]?.[style];
      table = null;
      if (stated) {
        const byPpem = new Map<number, ReadonlyMap<string, number>>();
        for (const [ppem, differences] of Object.entries(stated)) {
          const byCharacter = new Map<string, number>();
          for (const [dots, characters] of Object.entries(differences)) for (const character of characters) byCharacter.set(character, Number(dots));
          byPpem.set(Number(ppem), byCharacter);
        }
        table = byPpem;
      }
      hintedAdvanceTables.set(cached, table);
    }
    if (table) return table;
  }
  return null;
}

export function nativeTextAdvances(text: string, fontCss: string, dpi: number, scale = 1, purpose: "paint" | "measure" = "paint", advanceOffset = 0): Array<{ text: string; spacing: number; advance?: number }> {
  if (typeof document === "undefined" || /[\u0590-\u0fff]/u.test(text)) return [{ text, spacing: 0 }];
  if (textContext === undefined) {
    textContext = document.createElement("canvas").getContext("2d");
    document.fonts?.addEventListener("loadingdone", (event) => {
      if (!dropAdvancesTouched(event)) return;
      // Texts measured with what went are measured again (SpreadsheetNativeText).
      typographyGeneration += 1;
      for (const listener of [...typographyListeners]) listener();
    });
  }
  if (!textContext) return [{ text, spacing: 0 }];
  // Painted on the print page at the printer's advances (printerTextDpi);
  // the screen's grid is 96 dpi.
  // A text registered as designed (setNativeTextDesignedLines) is painted at
  // its design advances there (-1).
  const printer = purpose === "paint" && dpi === 96 && printerTextDpi ? (designedTextOwners.size && designedTextLine(text) ? -1 : printerTextDpi) : 0;
  const answerKey = `${text}\u0000${fontCss}\u0000${dpi}\u0000${scale}\u0000${advanceOffset}\u0000${printer}`;
  const known = advancesCache.get(answerKey);
  // A copy: callers own what they are given.
  if (known) return known.map((run) => ({ ...run }));
  const answer = measureNativeTextAdvances(textContext, text, fontCss, dpi, scale, advanceOffset, printer);
  if (advancesCache.size > 20000) advancesCache.clear();
  advancesCache.set(answerKey, answer.map((run) => ({ ...run })));
  return answer;
}

/**
 * One character's printed advance, whole printer dots, as Windows Excel sets
 * it (the API's writer the same, notebooks.py FaceMeasure.printed / unit):
 * its design advance (`units`, thousandths of an em) at the printer's pixels
 * per em, the half dot rounded up -- but
 * - a Hangul syllable or compatibility jamo a whole em in every face, as each
 *   Korean face of Windows sets it (Malgun Gothic, Gulim, Batang, ...: 92 dots
 *   at 11 pt, lookwin p187-lw), where the stand-ins set it 0.92 em;
 * - in a fixed-pitch East Asian face of GDI's (`cells`: ＭＳ ゴシック, 黑体,
 *   新宋体, 楷体, 굴림체, ...; excel-windows-cell-pitch-faces.json) an ASCII
 *   character one half-width cell, half the em rounded up to a whole dot,
 *   whatever stand-in draws it, and a character an em wide two cells: 76 dots
 *   at 9 pt where the em is 75, 118 at 14 pt (lookwin p186-lw; 宋体 and ＭＳ Ｐ
 *   ゴシック set theirs at the em) -- but the ideographic space at the em;
 * - `setShort`: the half dot rounded down from 51 dots an em.
 */
export function printedAdvanceDots(character: string, units: number, ppem: number, cells: boolean, setShort = false): number {
  const code = character.length <= 2 && [...character].length === 1 ? character.codePointAt(0) ?? 0 : 0;
  if ((code >= 0xac00 && code <= 0xd7a3) || (code >= 0x3131 && code <= 0x318e)) units = 1000;
  else if (cells && ((code >= 0x20 && code <= 0x7e) || code === 0x20a9)) units = 500;
  // The ideographic space is no cell's: it prints at the em (ＭＳ 明朝, ＭＳ
  // ゴシック, 新宋体 at 9 pt: 75 dots where their Han stand 76; xlsx backlog
  // bl5/r70 r70r).
  if (cells && Math.abs(units - 1000) < 0.5 && code !== 0x3000) return 2 * Math.floor(ppem / 2 + 0.5);
  const exact = units / 1000 * ppem;
  // `setShort`: a landscape print's pure-ASCII string in a GDI East Asian
  // face, as the printer sets it (setNativeTextLandscapePrint).
  if (setShort && ppem > 50 && Math.abs(exact - Math.floor(exact) - 0.5) < 1e-3) return Math.floor(exact);
  return Math.round(exact);
}

function measureNativeTextAdvances(context: CanvasRenderingContext2D, text: string, fontCss: string, dpi: number, scale: number, advanceOffset: number, printer = 0): Array<{ text: string; spacing: number; advance?: number }> {
  const textContext = context;
  const fontSize = fontCss.match(/([0-9]+(?:\.[0-9]+)?)(px|pt)/);
  if (!fontSize) return [{text,spacing:0}];
  // Screen-copy vectors position visible glyphs on the native device grid
  // even when the declared font is absent. Measure the resolved fallback
  // stack below; print-PDF text advances are not a screen positioning oracle.
  const pixels = Number(fontSize[1]) * (fontSize[2] === "pt" ? PX_PER_PT : 1);
  const safeScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  const paintedFont = fontCss.replace(fontSize[0], `${pixels * safeScale}px`);
  const quantum = nativeTextLayoutQuantum();
  let totalDesired = advanceOffset;
  let previousSnapped = quantum ? Math.round(advanceOffset / quantum) * quantum : advanceOffset;
  const idealFont = fontCss.replace(fontSize[0], "1000px");
  graphemes ??= new Intl.Segmenter(undefined, { granularity: "grapheme" });
  const segments = Array.from(graphemes.segment(text), ({ segment }) => segment);
  // The design widths from the served font files (lib/text-engine) where the
  // engine knows every grapheme's face: the same in every browser. Else the
  // browser's own, as before.
  const engineIdeals = spreadsheetEngineIdealAdvances(segments, fontCss);
  // The design and painted widths come from contexts kept in their fonts
  // (fontMeasureContext). The painted widths are measured, not scaled from
  // the design ones: hinted CJK faces are not linear (0.33 px at 12 px).
  const ideals = segments.map((segment, index) => {
    if (segment === "\n" || segment === "\r") return 0;
    // The glyph's design width, read once per face: every cell on screen
    // asks for the same few dozen characters on every render.
    const idealKey = `${idealFont}:${segment}`;
    let ideal = engineIdeals?.[index] ?? idealAdvanceCache.get(idealKey);
    if (ideal === undefined) {
      ideal = (fontMeasureContext(idealFont, "none") ?? textContext).measureText(segment).width;
      if (idealAdvanceCache.size > 12000) idealAdvanceCache.clear();
      idealAdvanceCache.set(idealKey, ideal);
    }
    return ideal;
  });
  const painted = segments.map((segment) => {
    if (segment === "\n" || segment === "\r") return 0;
    const key = `${paintedFont}:${segment}`;
    let width = advanceCache.get(key);
    if (width === undefined) {
      width = (fontMeasureContext(paintedFont, "none") ?? textContext).measureText(segment).width;
      if (advanceCache.size > 12000) advanceCache.clear();
      advanceCache.set(key, width);
    }
    return width;
  });
  // A fixed-pitch East Asian face of GDI's prints in cells (printedAdvanceDots).
  const primaryFamily = printer > 0 ? stackPrimaryFamily(fontCss.slice(fontCss.indexOf(fontSize[0]) + fontSize[0].length).replace(/^\s*\/\s*[^\s]+/, "")) : "";
  const cells = printer > 0 && excelWindowsCellPitch(primaryFamily);
  // A landscape print's pure-ASCII string in a GDI East Asian face: the
  // glyphs at the printer's own advances, the line as wide as Excel lays it
  // out (the dots the glyphs fall short stand after the line's last one).
  const setShort = printer > 0 && nativeTextLandscapePrint() && /^[\u0000-\u007f]*$/.test(text) && halfEmAsciiFace(primaryFamily);
  const lineEnds = new Set<number>();
  if (setShort) {
    let last = -1;
    segments.forEach((segment, index) => {
      if (segment === "\n" || segment === "\r") {
        if (last >= 0) lineEnds.add(last);
        last = -1;
      } else last = index;
    });
    if (last >= 0) lineEnds.add(last);
  }
  let shortDots = 0;
  // An old TrueType face's ASCII at GDI's hinted advances (officeHintedAdvanceDeltas).
  const stylePrefix = fontCss.slice(0, fontCss.indexOf(fontSize[0]));
  const hinted = printer > 0 && !cells
    ? officeHintedAdvanceDeltas(primaryFamily, /\b(?:bold|[6-9]00)\b/.test(stylePrefix), /\b(?:italic|oblique)\b/.test(stylePrefix))
    : null;
  return segments.map((segment, index) => {
    if (segment === "\n" || segment === "\r") {
      totalDesired = 0;
      previousSnapped = 0;
      return { text: segment, spacing: 0 };
    }
    const idealWidth = ideals[index] * pixels / 1000;
    // At the printer: whole pixels per em at the page's scale, whole dots an
    // advance (lookwin R223-X13: Arial Narrow 8 at 90 % rounds at 60 px per
    // em, as Windows Excel and the API print it, 169.2 pt a note line where
    // 67 px per em scaled to 90 % made 170.5).
    const ppem = Math.round(pixels * safeScale * printer / 96);
    let printedDots = printer > 0 ? printedAdvanceDots(segment, ideals[index], ppem, cells, setShort) + (hinted?.get(ppem)?.get(segment) ?? 0) : 0;
    if (setShort) {
      shortDots += printedAdvanceDots(segment, ideals[index], ppem, cells) + (hinted?.get(ppem)?.get(segment) ?? 0) - printedDots;
      if (lineEnds.has(index)) {
        printedDots += shortDots;
        shortDots = 0;
      }
    }
    const desired = printer > 0
      ? printedDots * 96 / printer
      : printer < 0
        ? idealWidth * safeScale
        : Math.round(idealWidth * dpi / 96) * 96 / dpi * safeScale;
    totalDesired += desired;
    const snapped = quantum ? Math.round(totalDesired / quantum) * quantum : totalDesired;
    const targetAdvance = snapped - previousSnapped;
    previousSnapped = snapped;
    const spacing = (targetAdvance - painted[index]) / safeScale;
    return { text: segment, spacing, advance: desired };
  });
}
