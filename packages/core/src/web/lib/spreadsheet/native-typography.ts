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
import { ensureOfficeFauxBoldProbe, officeFamilyHasNoBold } from "../office-fonts/faux-bold";
import { ensureOfficeFontFaces, loadOfficeFamilySpace, officeFontDeclared, officeFontSubstitutes, setOfficeEmboldenedRegularCuts } from "../office-fonts/substitutes";
import { columnBaseWidthToPixels, columnFileWidthToPixels } from "./grid-sizing";
import { isSpreadsheetFauxBoldStroke, setSpreadsheetFauxBoldPrinted, spreadsheetEngineIdealAdvances, spreadsheetEngineTextWidth, spreadsheetFauxBoldCss, spreadsheetTextEngineOn } from "./spreadsheet-text-engine";
import { withScriptFallbacks } from "../office-fonts/script-fallbacks";
import { OFFICE_COMPLEX_SCALE_PREFIX, officeComplexScaleStack } from "./office-complex-scales";
import { excelFontRecordOf, excelWindowsDrawnFaces, type ExcelFontRecordClass } from "../office-fonts/excel-windows-faces";
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
export function resolveNativeRowHeights(metrics: NativeRowMetrics | undefined, heights: Record<string,number> | undefined, defaultHeight: number | undefined, fontRowHeight?: number) {
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
  if (automatic && Math.abs(fontRowHeight - metrics.default.pixels) >= .5 && Math.abs(fontRowHeight - stated) >= .5) {
    for (const [key,row] of Object.entries(metrics.rows)) {
      if (row.customHeight || heights?.[key] !== row.pixels) continue;
      delete resolved[key];
    }
  }
  return {heights:resolved,defaultHeight:automatic ? clampNativeAutomaticRowHeight(fontRowHeight) : pairedDefault ? metrics.default.heightPt/.75 : defaultHeight};
}

/** Whether a default row the file marks custom (sheetFormatPr customHeight)
 * holds the rows that state no height of their own at it: only while no row
 * of the sheet states a custom height. Once one does, Windows Excel fits
 * those rows to their content (lookwin p97e-lw: a wrapped bold W92000004 in a
 * row with no ht stands one line, two once row 1, 3 or 5 states ht
 * customHeight; R346-X27). The API's writer (custom_row_stated) is the same. */
export function nativeDefaultHoldsRows(metrics: NativeRowMetrics | undefined): boolean {
  if (!metrics?.default.customHeight) return false;
  return !Object.values(metrics.rows).some((row) => row.customHeight);
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
export function nativeDefaultWidth(grid: NativeGrid, digitWidth: number, dpi: number) {
  return grid.defaultWidth === undefined
    ? columnBaseWidthToPixels(grid.baseColumnWidth, digitWidth, dpi) * 96 / dpi
    : nativeWidthToPixels(grid.defaultWidth, digitWidth, dpi);
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
    return servedStack(`"${face}", "Arial", ${arial}${officeComplexScaleStack(face)}"Microsoft YaHei", "Malgun Gothic", "Mangal", "Tahoma", sans-serif`);
  }
  ensureOfficeFontFaces(face);
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
  return servedStack(`${face ? `"${face}", ${complex}${substitutes}${scriptSerif}${hanSerif}` : ""}"Calibri","Microsoft YaHei", "Malgun Gothic", "Mangal", "Tahoma", "PingFang SC", sans-serif`);
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
    if (!GENERIC_FAMILIES.has(name.toLowerCase()) && !name.startsWith(OFFICE_COMPLEX_SCALE_PREFIX)) ensureOfficeFontFaces(name);
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
  if (bold && officeFamilyHasNoBold(family)) {
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
 * still break where the screen breaks them ("measure"); only the painted
 * advances follow the printer. The API's sheet PDF writer is the same
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
  return segments.map((segment, index) => {
    if (segment === "\n" || segment === "\r") {
      totalDesired = 0;
      previousSnapped = 0;
      return { text: segment, spacing: 0 };
    }
    const idealWidth = ideals[index] * pixels / 1000;
    // At the printer: whole pixels per em, whole dots an advance.
    const desired = printer > 0
      ? Math.round(ideals[index] / 1000 * Math.round(pixels * printer / 96)) * 96 / printer * safeScale
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
